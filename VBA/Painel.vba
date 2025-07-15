Sub Painel()

    Dim coligada As String, DataInicio As String, DataFim As String
    Dim usuario As String, senha As String
    Dim wsPainel As Worksheet
    Dim caminhoArquivo As String, nomeArquivo As String, caminhoCompleto As String
    Dim Anexo As String
    Dim g As Long, PriLin As Long, UltLin As Long
    Dim ws As Worksheet, tbl As ListObject
    Dim wbAnterior As Workbook, wbNovo As Workbook
    Dim pt As PivotTable
    Dim nomeAba As String

    ' Desativa cálculo automático para melhor performance
    Application.Calculation = xlCalculationManual

    ' Define a worksheet de controle
    Set wsPainel = ThisWorkbook.Sheets("Painel")
    
    ' Captura credenciais do painel
    usuario = wsPainel.Range("B6").Value
    senha = wsPainel.Range("B7").Value
    
    ' Define intervalo de linhas a percorrer
    PriLin = 15
    UltLin = 18

    With wsPainel
        For g = PriLin To UltLin
            If .Range("B" & g).Value = "Sim" Then
            
                coligada = .Range("A" & g).Value
                DataInicio = Format(.Range("C" & g).Value, "yyyy-MM-dd")
                caminhoArquivo = .Range("E" & g).Value
                nomeArquivo = .Range("F" & g).Value
                Anexo = .Range("H" & g).Value
                
                ' Validação de entrada
                If caminhoArquivo = "" Or nomeArquivo = "" Then
                    MsgBox "Caminho ou nome do arquivo ausente na linha " & g & ".", vbExclamation
                    GoTo Proximo
                End If

                ' Executa rotina principal com API
                Call Extrair_API_Nova(coligada, DataInicio, DataFim, usuario, senha)

                ' Verifica se deve gerar arquivo separado
                If Anexo = "Sim" Then
                
                    Set wbAnterior = ThisWorkbook
                    
                    On Error Resume Next
                    Set tbl = ws.ListObjects("Base_dados")
                    On Error GoTo 0
                    
                    nomeAba = "SALDO COL" & coligada
                    Debug.Print "Tentando copiar a aba: " & nomeAba
                    
                    ' Pequeno delay para garantir criação da aba
                    Application.Wait (Now + TimeValue("0:00:01"))
                    
                    ' Verifica existência da aba
                    On Error Resume Next
                    Set ws = ThisWorkbook.Worksheets(nomeAba)
                    On Error GoTo 0

                    If Not ws Is Nothing Then
                        ws.Copy
                        Set wbNovo = Workbooks(Workbooks.Count)

                        If Right(caminhoArquivo, 1) <> "\" Then
                            caminhoArquivo = caminhoArquivo & "\"
                        End If
                        caminhoCompleto = caminhoArquivo & nomeArquivo & ".xlsx"

                        Application.DisplayAlerts = False
                        wbNovo.SaveAs Filename:=caminhoCompleto, FileFormat:=xlOpenXMLWorkbook
                        
                        On Error Resume Next
                        For Each ws In wbNovo.Worksheets
                            For Each pt In ws.PivotTables
                                pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
                                pt.RefreshTable
                            Next pt
                        Next ws
                        On Error GoTo 0

                        wbNovo.Close SaveChanges:=False
                        Application.DisplayAlerts = True

                    Else
                        MsgBox "A aba " & nomeAba & " não foi encontrada."
                    End If
                    
Proximo:
                End If
            End If
        Next g
    End With

    Application.Calculation = xlCalculationAutomatic

End Sub
