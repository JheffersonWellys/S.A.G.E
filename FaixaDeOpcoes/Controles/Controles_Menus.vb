﻿Imports System.Drawing

Public Module Controles_Menus

#Region "GetImage"

    Public Function GetImage_GestaoDeAcessos() As Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_GestaoDeInfraestrutura() As Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_GestaoEducacional() As Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_GestaoAcademica() As Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_GestaoDeEventos() As Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_CalendarioAcademico_Exportacao() As Bitmap
        Return My.Resources.icn_Padrao
    End Function

    Public Function GetImage_MapaDeSala_Exportacao() As Bitmap
        Return My.Resources.icn_Padrao
    End Function

#End Region

#Region "GetEnabled"

    Public Function GetEnabled_GestaoDeAcessos() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GestaoDeInfraestrutura() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GestaoEducacional() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GestaoAcademica() As Boolean
        Return True
    End Function

    Public Function GetEnabled_GestaoDeEventos() As Boolean
        Return True
    End Function

    Public Function GetEnabled_CalendarioAcademico_Exportacao() As Boolean
        Return True
    End Function

    Public Function GetEnabled_MapaDeSala_Exportacao() As Boolean
        Return True
    End Function

#End Region

End Module
