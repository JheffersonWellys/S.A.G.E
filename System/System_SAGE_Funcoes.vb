Imports Microsoft.Office.Core

Public Module System_SAGE_Funcoes

#Region "Manipulação de Arquivos"

    Public Function SelecionarArquivo(titulo As String, extensoes As String, multiplaSelecao As Boolean) As String

        Dim app = Globals.ThisWorkbook.Application
        Dim fd As FileDialog = app.FileDialog(MsoFileDialogType.msoFileDialogFilePicker)

        With fd

            .Title = titulo
            .AllowMultiSelect = multiplaSelecao

            .Filters.Clear()

            Dim pattern As String = NormalizarExtensoes(extensoes)

            .Filters.Add("Arquivos (" & pattern & ")", pattern)

            If .Show = -1 Then

                If multiplaSelecao AndAlso .SelectedItems.Count > 0 Then

                    Return .SelectedItems.Item(1)

                ElseIf .SelectedItems.Count > 0 Then

                    Return .SelectedItems.Item(1)

                End If

            End If

        End With

        Return String.Empty

    End Function

    Public Function SelecionarArquivos(titulo As String, extensoes As String) As List(Of String)

        Dim lista As New List(Of String)
        Dim app = Globals.ThisWorkbook.Application
        Dim fd As FileDialog = app.FileDialog(MsoFileDialogType.msoFileDialogFilePicker)

        With fd

            .Title = titulo
            .AllowMultiSelect = True

            .Filters.Clear()

            Dim pattern As String = NormalizarExtensoes(extensoes)
            .Filters.Add("Arquivos (" & pattern & ")", pattern)

            If .Show = -1 Then

                For i As Integer = 1 To .SelectedItems.Count

                    lista.Add(.SelectedItems.Item(i))

                Next

            End If

        End With

        Return lista

    End Function

    Private Function NormalizarExtensoes(extensoes As String) As String

        Dim tokens = extensoes.Split(";"c) _
                              .Select(Function(t) t.Trim()) _
                              .Where(Function(t) t <> "")

        Dim pads = tokens.Select(Function(x)
                                     If x.StartsWith("*.") Then Return x
                                     If x.StartsWith(".") Then Return "*" & x
                                     If x.StartsWith("*") Then
                                         If x.Contains("."c) Then Return x
                                         Return "*." & x.TrimStart("*"c)
                                     End If
                                     Return "*." & x
                                 End Function)

        Return String.Join(";", pads)

    End Function

#End Region

#Region "Modo Foco - Aparência da Janela"

    Public Sub AplicarModoFoco()

        Dim wb = Globals.ThisWorkbook.Application

        Try
            wb.DisplayFormulaBar = False
            wb.ActiveWindow.DisplayHorizontalScrollBar = False
        Catch ex As Exception

        End Try

        For Each ws As Excel.Window In wb.Windows
            Try
                ws.DisplayHeadings = False
                ws.DisplayGridlines = False
                ws.DisplayWorkbookTabs = False
            Catch

            End Try
        Next
    End Sub

    Public Sub RemoverModoFoco()

        Dim wb = Globals.ThisWorkbook.Application

        Try
            wb.DisplayFormulaBar = True
            wb.ActiveWindow.DisplayHorizontalScrollBar = True
        Catch ex As Exception

        End Try

        For Each ws As Excel.Window In wb.Windows

            Try
                ws.DisplayHeadings = True
                ws.DisplayGridlines = True
                ws.DisplayWorkbookTabs = True
            Catch

            End Try

        Next

    End Sub

    Public Sub AlternarModoFoco(ativar As Boolean)
        If ativar Then
            AplicarModoFoco()
        Else
            RemoverModoFoco()
        End If
    End Sub

#End Region

End Module
