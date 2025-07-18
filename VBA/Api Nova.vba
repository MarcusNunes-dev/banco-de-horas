Sub Extrair_API_Nova(coligada As String, DataInicio As String, DataFim As String, login As String, Senha As String)

    Dim http As Object
    Dim url As String
    Dim response As String
    Dim linhas() As String
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim campos() As String
    Dim campoNome As String, campoValor As String
    Dim dict As Object
    Dim colunasEsperadas As Variant
    Dim colIndex As Long
    Dim tbl As ListObject
    Dim nomeAba As String
    Dim wsExistente As Worksheet
    Dim ultimaLinha As Long, ultimaColuna As Long

    ' === Definição das colunas esperadas no retorno da API ===
    colunasEsperadas = Array("COLIGADA", "CHAPA", "COLABORADOR", _
                             "PROJETO", "MAO DE OBRA", "SITUACAO", _
                             "DATA ADMISSAO", "DATA RESCISAO", "DESCR. CARGO", _
                             "SALDO_BH")

    ' === Captura de login e senha a partir de planilha ===
    login = Sheets("Painel").Range("B6").Value
    Senha = Sheets("Painel").Range("B7").Value

    ' === Montagem da URL da API ===
    url = "https://sua-api.com.br/api/framework/v1/consultaSQLServer/RealizaConsulta/ENDPOINT/0/A/?parameters=CODCOLIGADA=" & coligada & ";DATA_INICIO=" & DataInicio

    ' === Requisição HTTP ===
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Basic " & Base64Encode(login & ":" & Senha)
    http.setRequestHeader "Content-Type", "application/json"
    http.send

    response = http.responseText

    ' === Preparação da aba onde os dados serão exibidos ===
    nomeAba = "SALDO COL" & coligada

    ' Remove aba se já existir
    On Error Resume Next
    Set wsExistente = ThisWorkbook.Sheets(nomeAba)
    On Error GoTo 0

    If Not wsExistente Is Nothing Then
        Application.DisplayAlerts = False
        wsExistente.Delete
        Application.DisplayAlerts = True
    End If

    ' Cria nova aba
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nomeAba

    ' Preenche cabeçalho
    For i = 0 To UBound(colunasEsperadas)
        ws.Cells(1, i + 1).Value = colunasEsperadas(i)
    Next i

    ' Linha 2 temporária para permitir criação da tabela
    If ws.Cells(2, 1).Value = "" Then ws.Cells(2, 1).Value = "TEMP"
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = UBound(colunasEsperadas) + 1

    ' Criação da Tabela
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, ultimaColuna)), , xlYes)
    tbl.Name = "Base_dados"

    ' Remove a linha temporária
    If ws.Cells(2, 1).Value = "TEMP" Then
        tbl.DataBodyRange.Rows(1).Delete
    End If

    ' Limpeza de dados anteriores
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
    End If

    ' === Processamento manual do JSON ===
    response = Replace(response, "},{", "}§{")
    response = Replace(response, "[", "")
    response = Replace(response, "]", "")
    linhas = Split(response, "§")

    Dim matrizDados() As Variant
    ReDim matrizDados(1 To UBound(linhas) + 1, 1 To UBound(colunasEsperadas) + 1)

    For i = 0 To UBound(linhas)
        linhas(i) = Replace(linhas(i), "{", "")
        linhas(i) = Replace(linhas(i), "}", "")
        campos = Split(linhas(i), ",")

        Set dict = CreateObject("Scripting.Dictionary")

        For j = 0 To UBound(campos)
            If InStr(campos(j), ":") > 0 Then
                partes = Split(campos(j), ":")
                If UBound(partes) >= 1 Then
                    campoNome = Trim(Replace(partes(0), """", ""))
                    campoValor = ""
                    For k = 1 To UBound(partes)
                        campoValor = campoValor & partes(k)
                        If k < UBound(partes) Then campoValor = campoValor & ":"
                    Next k
                    campoValor = Replace(campoValor, """", "")
                    campoValor = Replace(campoValor, "'", "")
                    campoValor = Trim(campoValor)

                    ' Exemplo de formatação de campo específico
                    If campoNome = "SALDO BANCO DE HORAS" Then
                        On Error Resume Next
                        If IsNumeric(campoValor) Then
                            horas = Int(campoValor / 60)
                            minutos = campoValor Mod 60
                            campoValor = Format(horas, "00") & ":" & Format(minutos, "00")
                        End If
                        On Error GoTo 0
                    End If

                    dict(campoNome) = campoValor
                End If
            End If
        Next j

        ' Preencher matriz com base nas colunas desejadas
        For j = 0 To UBound(colunasEsperadas)
            If dict.exists(colunasEsperadas(j)) Then
                matrizDados(i + 1, j + 1) = dict(colunasEsperadas(j))
            Else
                matrizDados(i + 1, j + 1) = ""
            End If
        Next j
    Next i

    ' Colar dados na planilha
    ws.Range("A2").Resize(UBound(matrizDados), UBound(matrizDados, 2)).Value = matrizDados

    ' MsgBox "Consulta finalizada com sucesso!", vbInformation

End Sub
