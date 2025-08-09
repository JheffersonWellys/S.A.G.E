Public Module Nucleo_Variaveis

    Public UIRibbon_SAGE As Office.IRibbonUI

    Public TabAtual As String = "Tb_Logon"

    Public Informacao_BancoDeDados_Configurado As Boolean = False
    Public Informacao_BancoDeDados_Caminho As String = ""
    Public Informacao_BancoDeDados_Versao As String = ""
    Public Informacao_BancoDeDados_EmUso As Boolean = True
    Public Informacao_BancoDeDados_UsuarioLogado_Nome As String = "Admin"
    Public Informacao_BancoDeDados_UsuarioLogado_Email As String = ""

    Public Visualizacao_Cronograma_CalendarioAcademico As Boolean = True
    Public Visualizacao_Cronograma_MapaDeSala As Boolean = True

End Module
