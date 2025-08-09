Imports System.Drawing

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon_Menu
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("S.A.G.E.Ribbon_Menu.xml")
    End Function

#Region "Retornos de Chamada da Faixa de Opções"

#Region "Ribbon"

    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        UIRibbon_SAGE = Me.ribbon
    End Sub

#End Region

#Region "Tabs"

#Region "GetVisible"

    Public Function GetVisible_Tb_Logon(control As Office.IRibbonControl) As Boolean
        Return Controles_Abas.GetVisible_Logon()
    End Function

    Public Function GetVisible_Tb_MenuInicial(control As Office.IRibbonControl) As Boolean
        Return Controles_Abas.GetVisible_MenuInicial()
    End Function

    Public Function GetVisible_Tb_EdicaoDeCronograma_CalendarioAcademico(control As Office.IRibbonControl) As Boolean
        Return Controles_Abas.GetVisible_EdicaoDeCronograma_CalendarioAcademico()
    End Function

    Public Function GetVisible_Tb_EdicaoDeCronograma_MapaDeSala(control As Office.IRibbonControl) As Boolean
        Return Controles_Abas.GetVisible_EdicaoDeCronograma_MapaDeSala()
    End Function

#End Region

#End Region

#Region "Buttons"

#Region "GetImage"

#Region "Ribbon"

    Public Function GetImage_Bttn_Logon_Login_IniciarSessao(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_IniciarSessao()
    End Function

    Public Function GetImage_Bttn_Logon_Configuracao_Informacoes_UsuarioLogado_SolicitarAcessoViaTeams(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_SolicitarAcessoViaTeams()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Logout_FinalizarSessao(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_FinalizarSessao()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarUsuarios(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarUsuarios()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarAcessosDeUsuarios(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarAcessosDeUsuarios()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarUnidadesEducacionais(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarUnidadesEducacionais()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarBlocos(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarBlocos()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarAndares(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarAndares()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarSalas(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarSalas()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarDocentes(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarDocentes()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAutorizacoesParaLecionar(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarAutorizacoesParaLecionar()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAtestados(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarAtestados()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarAreasProfissionais(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarAreasProfissionais()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarCursos(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarCursos()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarUnidadesCurriculares(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarUnidadesCurriculares()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarFeriados(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarFeriados()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarRecessos(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarRecessos()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarDatasEventuaisPorUnidadeEducacional(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GerenciarDatasEventuaisPorUnidadeEducacional()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Cronogramas_GestaoDeCalendariosAcademicos(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GestaoDeCalendariosAcademicos()
    End Function

    Public Function GetImage_Bttn_MenuInicial_Cronogramas_GestaoDeMapasDeSala(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_GestaoDeMapasDeSala()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_VoltarPara_MenuInicial(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_VoltarParaMenuInicial()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_CriarComIA()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_RecriarComIA()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Edicao_EditarCronogramaManualmente(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_Edicao_EditarCronogramaManualmente()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasEventuaisPorTurma(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_GerenciarDatasEventuaisPorTurma()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeProjetosIntegradores(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_GerenciarDatasDeProjetosIntegradores()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeEstagiosProfissionaisSupervisionados(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_GerenciarDatasDeEstagiosProfissionaisSupervisionados()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_VisualizarErros()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_VisualizarCronograma()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_VisualizarCalendarioAcademico()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_ExportacaoEmPDF()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_CalendarioAcademico_ExportacaoEmXLSX()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_VoltarPara_MenuInicial(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_VoltarParaMenuInicial()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_CriarComIA()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_RecriarComIA()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_VisualizarErros()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_VisualizarCronograma()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_VisualizarMapaDeSala()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_ExportacaoEmPDF()
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_MapaDeSala_ExportacaoEmXLSX()
    End Function

#End Region

#Region "Backstage"

    Public Function GetImage_Bttn_InformacoesSobreOSistema_Configuracoes_ConfigurarBancoDeDados(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_Configuracoes_ConfigurarBancoDeDados()
    End Function

    Public Function GetImage_Bttn_InformacoesSobreOSistema_Configuracoes_BancoDeDados_ReconfigurarBancoDeDados(control As Office.IRibbonControl) As Bitmap
        Return Controles_Botoes.GetImage_Configuracoes_ReconfigurarBancoDeDados()
    End Function

#End Region

#End Region

#Region "GetEnabled"

#Region "Ribbon"

    Public Function GetEnabled_Bttn_Logon_Login_IniciarSessao(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_IniciarSessao()
    End Function

    Public Function GetEnabled_Bttn_Logon_Configuracao_Informacoes_UsuarioLogado_SolicitarAcessoViaTeams(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Logout_FinalizarSessao(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_FinalizarSessao()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarUsuarios(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarUsuarios()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarAcessosDeUsuarios(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarAcessosDeUsuarios()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarUnidadesEducacionais(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarUnidadesEducacionais()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarBlocos(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarBlocos()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarAndares(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarAndares()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarSalas(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarSalas()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarDocentes(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarDocentes()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAutorizacoesParaLecionar(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarAutorizacoesParaLecionar()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAtestados(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarAtestados()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarAreasProfissionais(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarAreasProfissionais()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarCursos(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarCursos()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarUnidadesCurriculares(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarUnidadesCurriculares()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarFeriados(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarFeriados()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarRecessos(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarRecessos()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarDatasEventuaisPorUnidadeEducacional(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GerenciarDatasEventuaisPorUnidadeEducacional()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Cronogramas_GestaoDeCalendariosAcademicos(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GestaoDeCalendariosAcademicos()
    End Function

    Public Function GetEnabled_Bttn_MenuInicial_Cronogramas_GestaoDeMapasDeSala(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_GestaoDeMapasDeSala()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_CriarComIA()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_RecriarComIA()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Edicao_EditarCronogramaManualmente(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_EditarCronogramaManualmente()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasEventuaisPorTurma(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_GerenciarDatasEventuaisPorTurma()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeProjetosIntegradores(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_GerenciarDatasDeProjetosIntegradores()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeEstagiosProfissionaisSupervisionados(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_GerenciarDatasDeEstagiosProfissionaisSupervisionados()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_VisualizarErros()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_VisualizarCronograma()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_VisualizarCalendarioAcademico()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_ExportacaoEmPDF()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_CalendarioAcademico_ExportacaoEmXLSX()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_MapaDeSala_CriarComIA()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_MapaDeSala_RecriarComIA()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_MapaDeSala_VisualizarErros()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_MapaDeSala_VisualizarCronograma()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_MapaDeSala_VisualizarMapaDeSala()
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_MapaDeSala_ExportacaoEmPDF
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetEnabled_MapaDeSala_ExportacaoEmXLSX
    End Function

#End Region

#End Region

#Region "GetVisible"

#Region "Ribbon"

    Public Function GetVisible_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetVisible_CalendarioAcademico_Cronograma_VisualizarCronograma
    End Function

    Public Function GetVisible_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetVisible_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico
    End Function

    Public Function GetVisible_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetVisible_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma
    End Function

    Public Function GetVisible_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala(control As Office.IRibbonControl) As Boolean
        Return Controles_Botoes.GetVisible_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala
    End Function

#End Region

#End Region

#Region "OnAction"

#Region "Ribbon"

    Public Sub OnAction_Bttn_Logon_Login_IniciarSessao(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_IniciarSessao()
    End Sub
    Public Sub OnAction_Bttn_Logon_Configuracao_Informacoes_UsuarioLogado_SolicitarAcessoViaTeams(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_SolicitarAcessoViaTeams()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Logout_FinalizarSessao(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_FinalizarSessao()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarUsuarios(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarUsuarios()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeAcessos_GerenciarAcessosDeUsuarios(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarAcessosDeUsuarios()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarUnidadesEducacionais(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarUnidadesEducacionais()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarBlocos(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarBlocos()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarAndares(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarAndares()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeInfraestrutura_GerenciarSalas(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarSalas()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarDocentes(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarDocentes()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAutorizacoesParaLecionar(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarAutorizacoesParaLecionar()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoEducacional_GerenciarAtestados(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarAtestados()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarAreasProfissionais(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarAreasProfissionais()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarCursos(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarCursos()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoAcademica_GerenciarUnidadesCurriculares(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarUnidadesCurriculares()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarFeriados(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarFeriados()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarRecessos(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarRecessos()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_GestaoDeEventos_GerenciarDatasEventuaisPorUnidadeEducacional(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GerenciarDatasEventuaisPorUnidadeEducacional()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Cronogramas_GestaoDeCalendariosAcademicos(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GestaoDeCalendariosAcademicos()
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Cronogramas_GestaoDeMapasDeSala(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_GestaoDeMapasDeSala()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_VoltarParaMenuInicial()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_CriarComIA()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_RecriarComIA()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Edicao_EditarCronogramaManualmente(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_Edicao_EditarCronogramaManualmente()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasEventuaisPorTurma(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_GerenciarDatasEventuaisPorTurma()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeProjetosIntegradores(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_GerenciarDatasDeProjetosIntegradores()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_GestaoDeEventos_GerenciarDatasDeEstagiosProfissionaisSupervisionados(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_GerenciarDatasDeEstagiosProfissionaisSupervisionados()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_Cronograma_VisualizarCronograma()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_Cronograma_VisualizarCalendarioAcademico()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_ExportacaoVisualizarErros()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_ExportacaoEmPDF()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_CalendarioAcademico_ExportacaoEmXLSX()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_VoltarParaMenuInicial()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_CriarComIA()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_RecriarComIA()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_ExportacaoVisualizarErros()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_VisualizarCronograma()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarMapaDeSala(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_VisualizarMapaDeSala()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_ExportacaoEmPDF()
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_MapaDeSala_ExportacaoEmXLSX()
    End Sub

#End Region

#Region "Backstage"

    Public Sub OnAction_Bttn_InformacoesSobreOSistema_Configuracoes_ConfigurarBancoDeDados(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_Configuracoes_ConfigurarBancoDeDados()
    End Sub

    Public Sub OnAction_Bttn_InformacoesSobreOSistema_Configuracoes_BancoDeDados_ReconfigurarBancoDeDados(control As Office.IRibbonControl)
        Call Acoes_Botoes.OnAction_Configuracoes_ReconfigurarBancoDeDados()
    End Sub

#End Region

#End Region

#End Region

#Region "Menus"

#Region "GetImage"

    Public Function GetImage_Mn_MenuInicial_Configuracoes_GestaoDeAcessos(control As Office.IRibbonControl) As Bitmap
        Return Controles_Menus.GetImage_GestaoDeAcessos
    End Function

    Public Function GetImage_Mn_MenuInicial_Configuracoes_GestaoDeInfraestrutura(control As Office.IRibbonControl) As Bitmap
        Return Controles_Menus.GetImage_GestaoDeInfraestrutura
    End Function

    Public Function GetImage_Mn_MenuInicial_Configuracoes_GestaoEducacional(control As Office.IRibbonControl) As Bitmap
        Return Controles_Menus.GetImage_GestaoEducacional
    End Function

    Public Function GetImage_Mn_MenuInicial_Configuracoes_GestaoAcademica(control As Office.IRibbonControl) As Bitmap
        Return Controles_Menus.GetImage_GestaoAcademica
    End Function

    Public Function GetImage_Mn_MenuInicial_Configuracoes_GestaoDeEventos(control As Office.IRibbonControl) As Bitmap
        Return Controles_Menus.GetImage_GestaoDeEventos
    End Function

    Public Function GetImage_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao(control As Office.IRibbonControl) As Bitmap
        Return Controles_Menus.GetImage_CalendarioAcademico_Exportacao
    End Function

    Public Function GetImage_Mn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao(control As Office.IRibbonControl) As Bitmap
        Return Controles_Menus.GetImage_MapaDeSala_Exportacao
    End Function

#End Region

#Region "GetEnabled"

    Public Function GetEnabled_Mn_MenuInicial_Configuracoes_GestaoDeAcessos(control As Office.IRibbonControl) As Boolean
        Return Controles_Menus.GetEnabled_GestaoDeAcessos
    End Function

    Public Function GetEnabled_Mn_MenuInicial_Configuracoes_GestaoDeInfraestrutura(control As Office.IRibbonControl) As Boolean
        Return Controles_Menus.GetEnabled_GestaoDeInfraestrutura
    End Function

    Public Function GetEnabled_Mn_MenuInicial_Configuracoes_GestaoEducacional(control As Office.IRibbonControl) As Boolean
        Return Controles_Menus.GetEnabled_GestaoEducacional
    End Function

    Public Function GetEnabled_Mn_MenuInicial_Configuracoes_GestaoAcademica(control As Office.IRibbonControl) As Boolean
        Return Controles_Menus.GetEnabled_GestaoAcademica
    End Function

    Public Function GetEnabled_Mn_MenuInicial_Configuracoes_GestaoDeEventos(control As Office.IRibbonControl) As Boolean
        Return Controles_Menus.GetEnabled_GestaoDeEventos
    End Function

    Public Function GetEnabled_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao(control As Office.IRibbonControl) As Boolean
        Return Controles_Menus.GetEnabled_CalendarioAcademico_Exportacao
    End Function

    Public Function GetEnabled_Mn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao(control As Office.IRibbonControl) As Boolean
        Return Controles_Menus.GetEnabled_MapaDeSala_Exportacao
    End Function

#End Region

#End Region

#Region "Groups"

#Region "GetVisible"

#Region "Ribbon"

    Public Function GetVisible_Grp_Logon_Login(control As Office.IRibbonControl) As Boolean
        Return GetVisible_Logon_Login()
    End Function

    Public Function GetVisible_Grp_Logon_Configuracao_Informacoes_BancoDeDadosNaoConfigurado(control As Office.IRibbonControl) As Boolean
        Return GetVisible_Informacoes_BancoDeDadosNaoConfigurado()
    End Function

    Public Function GetVisible_Grp_Logon_Configuracao_Informacoes_UsuarioLogado(control As Office.IRibbonControl) As Boolean
        Return GetVisible_Informacoes_UsuarioLogado()
    End Function

#End Region

#Region "Backstage"

    Public Function GetVisible_Grp_InformacoesSobreOSistema_Configuracoes_Alerta(control As Office.IRibbonControl) As Boolean
        Return GetVisible_InformacoesSobreOSistema_Configuracoes_Alerta()
    End Function

    Public Function GetVisible_Grp_InformacoesSobreOSistema_Configuracoes_BancoDeDados(control As Office.IRibbonControl) As Boolean
        Return GetVisible_InformacoesSobreOSistema_Configuracoes_BancoDeDados()
    End Function

#End Region

#End Region

#End Region

#Region "Labels"

#Region "Ribbon"

    Public Function GetLabel_LblCntrl_Logon_Configuracao_Informacoes_UsuarioLogado_Usuario(control As Office.IRibbonControl) As String
        Return $"O sistema está em uso por: [{Informacao_BancoDeDados_UsuarioLogado_Nome}]!"
    End Function

#End Region

#Region "Backstage"

    Public Function GetLabel_InformacoesSobreOSistema_Configuracoes_BancoDeDados_CaminhoBanco(control As Office.IRibbonControl) As String
        Return "Caminho: ..."
    End Function

    Public Function GetLabel_Lbl_InformacoesSobreOSistema_Configuracoes_BancoDeDados_VersaoBanco(control As Office.IRibbonControl) As String
        Return "Versão: ..."
    End Function

    Public Function GetLabel_Lbl_InformacoesSobreOSistema_Configuracoes_BancoDeDados_UltimaAtualizacao(control As Office.IRibbonControl) As String
        Return "Última Atualização: ..."
    End Function

#End Region

#End Region

#Region "Image Controls"

#Region "Backstage"

    Public Function GetImage_ImgCntrl_InformacoesSobreOSistema_Sistema_Logomarca(control As Office.IRibbonControl) As Bitmap
        Return GetImage_Sistema_Logomarca()
    End Function

#End Region

#End Region

#End Region

#Region "Auxiliares"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
