Public Class Clss_RegistroAcesso

    Public Property IdRegistroAcesso As Integer
    Public Property IdUsuario As Integer
    Public Property DataHoraEntrada As DateTime
    Public Property DataHoraSaida As DateTime?
    Public Property SessaoAtiva As Integer
    Public Property CriadoPor As Integer
    Public Property CriadoEm As DateTime
    Public Property AtualizadoPor As Integer?
    Public Property AtualizadoEm As DateTime?
    Public Property StatusRegistro As Integer

End Class