
<Runtime.InteropServices.ComVisible(True)> _
Public Class Rbbn_SAGE
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("S.A.G.E.Rbbn_SAGE.xml")
    End Function

#Region "Retornos de Chamada da Faixa de Opções"

#Region "Ribbon"

    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

#End Region

#Region "Tabs"

#Region "GetVisible"

    Public Function GetVisible_Tb_Logon(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_MenuInicial(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_GestaoDeAcesso(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_GestaoDeInfraestrutura(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_GestaoEducacional(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_GestaoAcademica(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_GestaoDeEventos(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_EdicaoDeCronograma_CalendarioAcademico(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Tb_EdicaoDeCronograma_MapaDeSala(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

#End Region

#End Region

#Region "Groups"

#Region "GetVisible"

    Public Function GetVisible_Grp_InformacoesSobreOSistema_Configuracoes_Alerta(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetVisible_Grp_InformacoesSobreOSistema_Configuracoes_BancoDeDados(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

#End Region

#End Region

#Region "Buttons"

#Region "GetImage"

    Public Function GetImage_Bttn_Logon_Login_IniciarSessao(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Logout_FinalizarSessao(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_IrPara_GestaoDeAcesso(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_IrPara_GestaoDeInfraestrutura(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_IrPara_GestaoEducacional(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_IrPara_GestaoAcademica(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Configuracoes_IrPara_GestaoDeEventos(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Cronogramas_IrPara_EdicaoDeCronograma_CalendarioAcademico(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_MenuInicial_Cronogramas_IrPara_EdicaoDeCronograma_MapaDeSala(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_VoltarPara_MenuInicial(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_VoltarPara_MenuInicial(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_VoltarPara_MenuInicial(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_VoltarPara_MenuInicial(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_VoltarPara_MenuInicial(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_VoltarPara_MenuInicial(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Edicao_EditarCronogramaManualmente(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Adicionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Editar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Remover(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Visualizar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Restaurar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_VoltarPara_MenuInicial(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_InformacoesSobreOSistema_Configuracoes_ConfigurarBancoDeDados(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Bttn_InformacoesSobreOSistema_Configuracoes_BancoDeDados_ReconfigurarBancoDeDados(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

#End Region

#Region "GetEnabled"
    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Edicao_EditarCronogramaManualmente(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Adicionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Editar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Remover(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Visualizar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Restaurar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_VisualizarCronograma(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

#End Region

#Region "OnAction"

    Public Sub OnAction_Bttn_Logon_Login_IniciarSessao(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Logout_FinalizarSessao(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_IrPara_GestaoDeAcesso(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_IrPara_GestaoDeInfraestrutura(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_IrPara_GestaoEducacional(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_IrPara_GestaoAcademica(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Configuracoes_IrPara_GestaoDeEventos(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Cronogramas_IrPara_EdicaoDeCronograma_CalendarioAcademico(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_MenuInicial_Cronogramas_IrPara_EdicaoDeCronograma_MapaDeSala(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Docentes_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoEducacional_Configuracoes_Gerenciar_Atestados_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_Cursos_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Edicao_EditarCronogramaManualmente(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Adicionar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Editar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Remover(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Visualizar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados_Restaurar(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_VoltarPara_MenuInicial(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_CriarComIA(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Criacao_RecriarComIA(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Validacao_VisualizarErros(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Visualizacao_VisualizarCronograma(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmPDF(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao_EmXLSX(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_InformacoesSobreOSistema_Configuracoes_ConfigurarBancoDeDados(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

    Public Sub OnAction_Bttn_InformacoesSobreOSistema_Configuracoes_BancoDeDados_ReconfigurarBancoDeDados(control As Office.IRibbonControl)
        MsgBox("Ação executada para o botão: " & control.Id, vbInformation, "Ribbon S.A.G.E.")
    End Sub

#End Region

#End Region

#Region "Menus"

#Region "GetImage"

    Public Function GetImage_Mn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoEducacional_Configuracoes_Gerenciar_Docentes(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoEducacional_Configuracoes_Gerenciar_Atestados(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoAcademica_Configuracoes_Gerenciar_Cursos(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_Mn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao(control As Office.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.icn_Padrao
    End Function

#End Region

#Region "GetEnabled"

    Public Function GetEnabled_Mn_GestaoDeAcesso_Configuracoes_Gerenciar_Usuarios(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeAcesso_Configuracoes_Gerenciar_AcessosDeUsuario(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_UnidadesEducacionais(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Blocos(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_Andares(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeInfraestrutura_Configuracoes_Gerenciar_SalasDeAula(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoEducacional_Configuracoes_Gerenciar_Docentes(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoEducacional_Configuracoes_Gerenciar_AutorizacoesParaLecionar(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoEducacional_Configuracoes_Gerenciar_Atestados(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoAcademica_Configuracoes_Gerenciar_AreasProfissionais(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoAcademica_Configuracoes_Gerenciar_Cursos(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoAcademica_Configuracoes_Gerenciar_UnidadesCurriculares(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeEventos_Configuracoes_Gerenciar_Feriados(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeEventos_Configuracoes_Gerenciar_Recessos(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_GestaoDeEventos_Configuracoes_Gerenciar_DatasEventuais(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_DatasEventuais(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_ProjetosIntegradores(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Configuracao_EstagiosProfissionaisSupervisionados(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_EdicaoDeCronograma_CalendarioAcademico_Cronograma_Exportacao(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

    Public Function GetEnabled_Mn_EdicaoDeCronograma_MapaDeSala_Cronograma_Exportacao(control As Office.IRibbonControl) As Boolean
        Return True
    End Function

#End Region

#End Region

#Region "Labels"

#Region "GetLabel"

    Public Function GetLabel_InformacoesSobreOSistema_Configuracoes_BancoDeDados_CaminhoBanco(control As Office.IRibbonControl) As String
        Return "Ribbon S.A.G.E"
    End Function

    Public Function GetLabel_Lbl_InformacoesSobreOSistema_Configuracoes_BancoDeDados_VersaoBanco(control As Office.IRibbonControl) As String
        Return "Ribbon S.A.G.E"
    End Function

    Public Function GetLabel_Lbl_InformacoesSobreOSistema_Configuracoes_BancoDeDados_UltimaAtualizacao(control As Office.IRibbonControl) As String
        Return "Ribbon S.A.G.E"
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
