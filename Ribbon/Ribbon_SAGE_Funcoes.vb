Public Module Ribbon_SAGE_Funcoes

#Region "Inicialização"

    Public Sub InicializarSistema()

        If String.IsNullOrEmpty(Globals.Planilha2.CONFIGURACAO_SISTEMA__CAMINHO_BANCO_DE_DADOS.Value2) Then

            RibbonUI_TabAtiva = "Tb_Configuracao"

        Else

            RibbonUI_TabAtiva = "Tb_Logon"

        End If

        Globals.Planilha1.Activate()

    End Sub

#End Region

#Region "Validação da Configuração do Banco"

    Public Function Action_VerificarSeBancoDeDadosEstaConfigurador() As Boolean

        Dim Status As Boolean = False

        If String.IsNullOrEmpty(Globals.Planilha2.CONFIGURACAO_SISTEMA__CAMINHO_BANCO_DE_DADOS.Value2) Then

            RibbonUI_TabAtiva = "Tb_Configuracao"
            Status = False

        Else

            RibbonUI_TabAtiva = "Tb_Logon"
            Status = True

        End If

        RibbonUI_SAGE.Invalidate()

        Return Status

    End Function

#End Region

#Region "Ações da Ribbon"

#Region "Ações Genéricas"

    Public Sub Action_AbrirAba(TabSelecionada As String)

        If Action_VerificarSeBancoDeDadosEstaConfigurador() = False Then
            MessageBox.Show("Sua sessão foi encerrada por conta de falta de conexão com o banco de dados!", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Globals.Planilha1.Activate()
            Exit Sub
        End If

        RibbonUI_TabAtiva = TabSelecionada
        RibbonUI_SAGE.Invalidate()

    End Sub

#End Region

#Region "Ações Específicas"

    Public Sub Action_ConfigurarBancoDeDados()

        Globals.Planilha2.CONFIGURACAO_SISTEMA__CAMINHO_BANCO_DE_DADOS.Value2 = SelecionarArquivo("Selecione o arquivo do banco de dados", ".db;sqlite;*.db3", False)
        Action_VerificarSeBancoDeDadosEstaConfigurador()

    End Sub

    Public Sub Action_ReconfigurarBancoDeDados()

        If String.IsNullOrEmpty(Globals.Planilha2.CONFIGURACAO_SISTEMA__CAMINHO_BANCO_DE_DADOS.Value2) And
                MessageBox.Show("Você deseja realmente selecionar outro banco de dados?", "Verificação!",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            Globals.Planilha2.CONFIGURACAO_SISTEMA__CAMINHO_BANCO_DE_DADOS.Value2 = SelecionarArquivo("Selecione o arquivo do banco de dados", ".db;sqlite;*.db3", False)
            Action_VerificarSeBancoDeDadosEstaConfigurador()

        End If

    End Sub

#End Region

#End Region

End Module
