
Public Module Controles_Grupos

#Region "GetVisible"

#Region "Ribbon"

    Public Function GetVisible_Logon_Login() As Boolean
        Return Informacao_BancoDeDados_Configurado = True And Informacao_BancoDeDados_EmUso = False
    End Function

    Public Function GetVisible_Informacoes_BancoDeDadosNaoConfigurado() As Boolean
        Return Informacao_BancoDeDados_Configurado = False
    End Function

    Public Function GetVisible_Informacoes_UsuarioLogado() As Boolean
        Return Informacao_BancoDeDados_EmUso = True And Informacao_BancoDeDados_Configurado = True
    End Function

#End Region

#Region "Backstage"

    Public Function GetVisible_InformacoesSobreOSistema_Configuracoes_Alerta() As Boolean
        Return Informacao_BancoDeDados_Configurado = False
    End Function

    Public Function GetVisible_InformacoesSobreOSistema_Configuracoes_BancoDeDados() As Boolean
        Return Informacao_BancoDeDados_Configurado = True
    End Function

#End Region

#End Region

End Module
