Imports System.Drawing

Public Module Controles_Abas

#Region "GetVisible"

    Public Function GetVisible_Logon() As Boolean
        Return TabAtual = "Tb_Logon"
    End Function

    Public Function GetVisible_MenuInicial() As Boolean
        Return TabAtual = "Tb_MenuInicial"
    End Function

    Public Function GetVisible_EdicaoDeCronograma_CalendarioAcademico() As Boolean
        Return TabAtual = "Tb_EdicaoDeCronograma_CalendarioAcademico"
    End Function

    Public Function GetVisible_EdicaoDeCronograma_MapaDeSala() As Boolean
        Return TabAtual = "Tb_EdicaoDeCronograma_MapaDeSala"
    End Function

#End Region

End Module
